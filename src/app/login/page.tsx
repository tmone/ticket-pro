"use client";

import * as React from "react";
import { useRouter } from "next/navigation";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Form, FormControl, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import { TicketCheck, LogIn } from "lucide-react";

const loginSchema = z.object({
  password: z.string().min(1, { message: "Password is required." }),
});

const HARDCODED_PASSWORD = "eventstaff"; // In a real app, this would be handled securely.

export default function LoginPage() {
  const [error, setError] = React.useState<string | null>(null);
  const router = useRouter();

  const form = useForm<z.infer<typeof loginSchema>>({
    resolver: zodResolver(loginSchema),
    defaultValues: {
      password: "",
    },
  });

  const onSubmit = (data: z.infer<typeof loginSchema>) => {
    if (data.password === HARDCODED_PASSWORD) {
      setError(null);
      sessionStorage.setItem("isAuthenticated", "true");
      router.push("/");
    } else {
      setError("Invalid password. Please try again.");
    }
  };
  
  return (
    <main className="flex min-h-screen flex-col items-center justify-center p-4">
      <Card className="w-full max-w-sm">
        <CardHeader className="text-center">
          <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center rounded-full bg-primary/10">
            <TicketCheck className="h-8 w-8 text-primary" />
          </div>
          <CardTitle className="text-2xl">TicketCheck Pro</CardTitle>
          <CardDescription>Please enter your password to access the dashboard.</CardDescription>
        </CardHeader>
        <CardContent>
          <Form {...form}>
            <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-4">
              <FormField
                control={form.control}
                name="password"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>Password</FormLabel>
                    <FormControl>
                      <Input type="password" placeholder="••••••••" {...field} />
                    </FormControl>
                    <FormMessage />
                  </FormItem>
                )}
              />
              {error && (
                <p className="text-sm font-medium text-destructive">{error}</p>
              )}
              <Button type="submit" className="w-full">
                <LogIn className="mr-2 h-4 w-4" />
                Log In
              </Button>
            </form>
          </Form>
        </CardContent>
      </Card>
    </main>
  );
}
